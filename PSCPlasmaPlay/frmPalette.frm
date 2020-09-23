VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPaletteMaker 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6750
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   6840
   ForeColor       =   &H00000000&
   Icon            =   "frmPalette.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   Visible         =   0   'False
   Begin VB.PictureBox picTEST 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   225
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   42
      Top             =   5805
      Width           =   3840
   End
   Begin VB.CommandButton cmdFileOps 
      Caption         =   "Open palette (PAL)"
      Height          =   330
      Index           =   2
      Left            =   4425
      TabIndex        =   41
      ToolTipText     =   " Does a 9-point average "
      Top             =   1464
      Width           =   1905
   End
   Begin VB.CommandButton cmdUSE 
      Caption         =   "Use this palette  v   - ->"
      Height          =   330
      Left            =   228
      TabIndex        =   39
      Top             =   4284
      Width           =   2565
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default  (Black)"
      Height          =   330
      Left            =   4410
      TabIndex        =   31
      Top             =   4032
      Width           =   2205
   End
   Begin VB.CommandButton cmdFileOps 
      Caption         =   "Save BPS"
      Height          =   330
      Index           =   1
      Left            =   4425
      TabIndex        =   30
      Top             =   1044
      Width           =   1905
   End
   Begin VB.CommandButton cmdFileOps 
      Caption         =   "Save palette (PAL)"
      Height          =   330
      Index           =   3
      Left            =   4425
      TabIndex        =   29
      Top             =   1872
      Width           =   1905
   End
   Begin VB.CommandButton cmdFileOps 
      Caption         =   "Open Button file (BPS)"
      Height          =   330
      Index           =   0
      Left            =   4425
      TabIndex        =   28
      Top             =   624
      Width           =   1905
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5805
      Top             =   5835
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "CD"
   End
   Begin VB.PictureBox PICPAL 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   225
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   8
      Top             =   4665
      Width           =   3840
   End
   Begin VB.Frame Frame1 
      Caption         =   "Color ----- Number of buttons"
      Height          =   1410
      Left            =   4410
      TabIndex        =   4
      Top             =   2460
      Width           =   2205
      Begin VB.HScrollBar HSNButtons 
         Height          =   225
         Index           =   2
         Left            =   1530
         Max             =   4
         Min             =   1
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   945
         Value           =   1
         Width           =   540
      End
      Begin VB.HScrollBar HSNButtons 
         Height          =   225
         Index           =   1
         Left            =   1530
         Max             =   4
         Min             =   1
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   615
         Value           =   1
         Width           =   540
      End
      Begin VB.HScrollBar HSNButtons 
         Height          =   225
         Index           =   0
         Left            =   1530
         Max             =   4
         Min             =   1
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   285
         Value           =   1
         Width           =   540
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1005
         Left            =   120
         ScaleHeight     =   1005
         ScaleWidth      =   900
         TabIndex        =   21
         Top             =   285
         Width           =   900
         Begin VB.OptionButton optCul 
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   24
            Top             =   15
            Width           =   300
         End
         Begin VB.OptionButton optCul 
            ForeColor       =   &H00008000&
            Height          =   210
            Index           =   1
            Left            =   45
            TabIndex        =   23
            Top             =   360
            Width           =   315
         End
         Begin VB.OptionButton optCul 
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   2
            Left            =   45
            TabIndex        =   22
            Top             =   720
            Width           =   315
         End
         Begin VB.Label LabC 
            Caption         =   "Blue"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   375
            TabIndex        =   27
            Top             =   690
            Width           =   465
         End
         Begin VB.Label LabC 
            Caption         =   "Green"
            ForeColor       =   &H00008000&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   26
            Top             =   345
            Width           =   465
         End
         Begin VB.Label LabC 
            Caption         =   "Red"
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   0
            Left            =   375
            TabIndex        =   25
            Top             =   15
            Width           =   435
         End
      End
      Begin VB.Label LabNB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   1095
         TabIndex        =   37
         Top             =   930
         Width           =   345
      End
      Begin VB.Label LabNB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   1
         Left            =   1095
         TabIndex        =   36
         Top             =   600
         Width           =   345
      End
      Begin VB.Label LabNB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   1095
         TabIndex        =   35
         Top             =   270
         Width           =   345
      End
   End
   Begin VB.PictureBox PIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4020
      Left            =   225
      Picture         =   "frmPalette.frx":0442
      ScaleHeight     =   268
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   268
      TabIndex        =   0
      Top             =   180
      Width           =   4020
      Begin VB.PictureBox PR 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   8
         Left            =   3840
         MousePointer    =   10  'Up Arrow
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   20
         Top             =   3855
         Width           =   150
      End
      Begin VB.PictureBox PR 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   7
         Left            =   3360
         MousePointer    =   10  'Up Arrow
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   19
         Top             =   3840
         Width           =   150
      End
      Begin VB.PictureBox PR 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   6
         Left            =   2895
         MousePointer    =   10  'Up Arrow
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   18
         Top             =   3855
         Width           =   150
      End
      Begin VB.PictureBox PR 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   5
         Left            =   2415
         MousePointer    =   10  'Up Arrow
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   17
         Top             =   3840
         Width           =   150
      End
      Begin VB.PictureBox PR 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   4
         Left            =   1935
         MousePointer    =   10  'Up Arrow
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   16
         Top             =   3855
         Width           =   150
      End
      Begin VB.PictureBox PR 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   3
         Left            =   1455
         MousePointer    =   10  'Up Arrow
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   15
         Top             =   3855
         Width           =   150
      End
      Begin VB.PictureBox PR 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   0
         Left            =   0
         MousePointer    =   10  'Up Arrow
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   1
         Top             =   3855
         Width           =   150
      End
      Begin VB.PictureBox PR 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   2
         Left            =   975
         MousePointer    =   10  'Up Arrow
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   3
         Top             =   3855
         Width           =   150
      End
      Begin VB.PictureBox PR 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   1
         Left            =   480
         MousePointer    =   10  'Up Arrow
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   2
         Top             =   3855
         Width           =   150
      End
   End
   Begin VB.Label Label1 
      Caption         =   " ^ Loaded PAL file palette"
      Height          =   225
      Index           =   2
      Left            =   240
      TabIndex        =   43
      Top             =   6270
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "<- Move the buttons"
      Height          =   228
      Index           =   0
      Left            =   4452
      TabIndex        =   40
      Top             =   216
      Width           =   1548
   End
   Begin VB.Label Label1 
      Caption         =   "^ RGB values (0-255)"
      Height          =   225
      Index           =   1
      Left            =   240
      TabIndex        =   38
      Top             =   5505
      Width           =   1560
   End
   Begin VB.Label LabR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R8"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   4110
      TabIndex        =   14
      Top             =   5205
      Width           =   330
   End
   Begin VB.Label LabR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R7"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   3615
      TabIndex        =   13
      Top             =   5160
      Width           =   330
   End
   Begin VB.Label LabR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R6"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   3135
      TabIndex        =   12
      Top             =   5145
      Width           =   330
   End
   Begin VB.Label LabR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R5"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   2595
      TabIndex        =   11
      Top             =   5115
      Width           =   330
   End
   Begin VB.Label LabR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R4"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   2190
      TabIndex        =   10
      Top             =   5145
      Width           =   330
   End
   Begin VB.Label LabR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R3"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1710
      TabIndex        =   9
      Top             =   5190
      Width           =   330
   End
   Begin VB.Label LabR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1260
      TabIndex        =   7
      Top             =   5145
      Width           =   330
   End
   Begin VB.Label LabR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R1"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   5160
      Width           =   330
   End
   Begin VB.Label LabR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   5145
      Width           =   330
   End
End
Attribute VB_Name = "frmPaletteMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmPaletteMaker by Robert Rayment

Option Explicit

'' To hold palette
' Public
' PalRGB() As Long

' For moving buttons
Private XS As Single, YS As Single
' Button positions
Private NumButtons() As Long
Private LX() As Long
Private LY() As Long
' Prev for Error
Private pNumButtons() As Long
Private pLX() As Long
Private pLY() As Long


Private ButtonCul() As Long
Private SCUL As Long    ' Selected color 0 R, 1 G, 2 B
' Button icons RArr.ico, GArr.ico, BArr.ico - could be redesigned.
Private aICOS As Boolean
Private AppPath$, FileSpec$

Private Sub cmdDefault_Click()
Dim j As Long
   INIT
   For j = 2 To 0 Step -1
      SCUL = j   ' Selected color
      HSNButtons(CInt(j)).Value = 2  ' Make 3 buttons
      Call HSNButtons_Change(CInt(j))
      RGBChanger
   Next j
   optCul(0) = True
   picTEST.BackColor = 0
   Caption = "  Palette maker"
End Sub

Private Sub cmdUSE_Click()
Dim N As Long

   For N = 0 To 255
      Colors(N) = RGB(PalRGB(0, N), PalRGB(1, N), PalRGB(2, N))
   Next N
   ' Expand palette to 512 colors
   For N = 255 To 1 Step -1
      Colors(2 * N - 1) = Colors(N)
   Next N
   Colors(510) = Colors(509)
   Colors(511) = Colors(510)

   For N = 2 To 510 Step 2
      Colors(N) = Colors(N - 1)
   Next N
   Colors(511) = Colors(510)
   
      ' Show palette
   For N = 0 To 511 Step 2
      Form1.PICPAL.Line (N \ 2, 0)-(N \ 2, Form1.PICPAL.Height), Colors(N)
   Next N
   Form1.PICPAL.Refresh
   
   Dim ix As Long, iy As Long
   For iy = 1 To 256 'sizey
   For ix = 1 To 256 'sizex
      ColArray(ix, iy) = Colors(IndexArray(ix, iy))
   Next ix
   Next iy
   Form1.LabPalName = LCase$(GetFileName(FileSpec$))

   Me.Hide
   Form1.Show
End Sub

'Const MaxNumButtons = 9 ' 2,3,5,9


Private Sub Form_Initialize()
'   m_hMod = LoadLibrary("shell32.dll")
'   InitCommonControls


   ReDim ButtonCul(0 To 2)
   ButtonCul(0) = vbRed
   ButtonCul(1) = &H8000&     ' green
   ButtonCul(2) = &HFF0000    ' blue
   
   INIT
End Sub

Private Sub INIT()
' Called from Form_Initialize  &  LoadBPS error
Dim nRGB As Long, px As Long
   ReDim NumButtons(0 To 2)   ' #R,#G,#B
   ReDim LY(0 To 2, 0 To 8)   ' Cul 0,1,2, ypos button 0,1,2
   ReDim LX(0 To 2, 0 To 8)   ' Cul 0,1,2, xpos button 0,1,2
   ' Prev Save for Error
   ReDim pNumButtons(0 To 2)   ' #R,#G,#B
   ReDim pLY(0 To 2, 0 To 8)   ' Cul 0,1,2, ypos button 0,1,2
   ReDim pLX(0 To 2, 0 To 8)   ' Cul 0,1,2, xpos button 0,1,2
   
   NumButtons(0) = 3          ' #R
   NumButtons(1) = 3          ' #G
   NumButtons(2) = 3          ' #B
   ReDim PalRGB(0 To 2, 0 To 255)   ' Palette - 0, Black
   
   ' x,y default pos & inverted RGB values ie 0 -> 255
   For px = 0 To 8
   For nRGB = 0 To 2
      If px = 8 Then
         LX(nRGB, px) = 255
      Else
         LX(nRGB, px) = px * 32
      End If
      LY(nRGB, px) = 255
   Next nRGB
   Next px
End Sub

Private Sub Form_Load()
Dim k As Long
   
   AppPath$ = App.Path
   If Right$(AppPath$, 1) <> "\" Then AppPath$ = AppPath$ & "\"
   FileSpec$ = AppPath$
   ' Test if icons there, avoids using a RES file
   aICOS = FileExists(AppPath$ & "RArr.ico")
   If aICOS Then aICOS = FileExists(AppPath$ & "GArr.ico")
   If aICOS Then aICOS = FileExists(AppPath$ & "BArr.ico")
   
   Show
   'BackColor = &HFFDBA1
   ' Size & position buttons
   PR(0).Top = 255
   If aICOS Then PR(0).Picture = LoadPicture(AppPath$ & "RArr.ico")
   For k = 0 To 8
      If aICOS Then
         PR(k).BorderStyle = 0
         PR(k).BackColor = vbWhite
         PR(k).Picture = PR(0).Picture
      Else
         PR(k).BorderStyle = 1
         PR(k).BackColor = vbRed
      End If
      PR(k).Height = 12
      PR(k).Width = 12
      PR(k).Top = PR(0).Top
      PR(k).Left = LX(0, k)
   Next k
   
   ' PIC size & location for BPS
   PIC.Height = 256 + PR(0).Height  ' 268
   PIC.Width = 256 + PR(0).Width    ' 268
   
   cmdUSE.Top = PIC.Top + PIC.Height + 10
   
   PICPAL.Width = 256
   PICPAL.Top = cmdUSE.Top + cmdUSE.Height + 10
   
   ' PIC & PICPAL Borders
   Line (PIC.Left - 2, PIC.Top - 2)-(PIC.Left + PIC.Width, PIC.Top + PIC.Height), 0, B
   Line (PICPAL.Left - 2, PICPAL.Top - 2)-(PICPAL.Left + PICPAL.Width + 1, PICPAL.Top + PICPAL.Height + 1), 0, B
   
   ' RGB labs
   LabR(0).Top = PICPAL.Top + PICPAL.Height + 10
   LabR(0).Left = PIC.Left
   For k = 0 To 8
      LabR(k) = 0
      LabR(k).BackColor = vbWhite
      LabR(k).Top = LabR(0).Top
      LabR(k).Left = LabR(0).Left + k * 32
   Next k
   Label1(1).Top = LabR(0).Top + Label1(0).Height + 10
   ' picTest size & location for PAL
   picTEST.Left = PICPAL.Left
   picTEST.Width = 256
   picTEST.Top = Label1(1).Top + Label1(1).Height + 10
   ' picTest Border
   Line (picTEST.Left - 2, picTEST.Top - 2)-(picTEST.Left + picTEST.Width + 1, picTEST.Top + picTEST.Height + 1), 0, B
   
   Label1(2).Top = picTEST.Top + picTEST.Height + 10
   ' NumButtons Labs Default = 3
   LabNB(0) = "3"
   LabNB(1) = "3"
   LabNB(2) = "3"
   SCUL = 0 ' Default Red
   optCul(0).Value = True
   HSNButtons(0) = 2 ' Gives 3 buttons
   HSNButtons(1) = 2
   HSNButtons(2) = 2
   ShowPAL
   Caption = "  Palette maker"
   frmPaletteMaker.Left = 60
   frmPaletteMaker.Top = 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Me.Hide
   Form1.Show
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'   FreeLibrary m_hMod
'
'End Sub
Private Sub HSNButtons_Change(Index As Integer)
' Index is color 0,1,2  Red,Green,Blue
Dim k As Long ', j As Long
   k = HSNButtons(Index).Value
   NumButtons(Index) = 2 ^ (k - 1) + 1  ' k=1,2,3,4 -> 2,3,5,9
   LabNB(Index) = NumButtons(Index)
   If Index = SCUL Then
      NewPalFromButtons Index
   End If
End Sub

Private Sub LabC_Click(Index As Integer)
' Re-direct to optCul
   optCul(Index) = True
End Sub


Private Sub optCul_Click(Index As Integer)
' Private SCUL As Long
' Index is color 0,1,2 Red,Green,Blue
   SCUL = Index   ' Selected color
   RGBChanger
End Sub

Private Sub RGBChanger()
' Called from optCul_Click  &  LoadBPS
Dim k As Long
   If aICOS Then
      Select Case SCUL
      Case 0: PR(0).Picture = LoadPicture(AppPath$ & "RArr.ico")
      Case 1: PR(0).Picture = LoadPicture(AppPath$ & "GArr.ico")
      Case 2: PR(0).Picture = LoadPicture(AppPath$ & "BArr.ico")
      End Select
      For k = 1 To 8
         PR(k).Picture = PR(0).Picture
      Next k
   Else
      For k = 0 To 8
         PR(k).BackColor = ButtonCul(SCUL)
      Next k
   End If
   
   NewPalFromButtons CInt(SCUL)
   For k = 0 To 8  ' Will skip Invis buttons
      PR(k).Top = LY(SCUL, k)
      PR(k).Left = LX(SCUL, k)
      LabR(k) = 255 - LY(SCUL, k)
   Next k
   'NewPalFromButtons CInt(SCUL)
End Sub

Private Sub NewPalFromButtons(TheIndex As Integer)
' Index is color 0,1,2  Red,Green,Blue
Dim SStep As Long
   ButtonChange NumButtons(TheIndex)
   SStep = 8 \ (NumButtons(TheIndex) - 1)  '2,3,5,9 -> 8,4,2,1
   DrawLines SStep
   GetRGBS SStep
   ShowPAL
End Sub

Private Sub ButtonChange(NButtons As Long)
Dim SStep As Long
Dim j As Long
   For j = 0 To 8
      LabR(j).Visible = False
      PR(j).Visible = False
   Next j
   SStep = 8 \ (NButtons - 1)  '2,3,5,9 -> 8,4,2,1
   For j = 0 To 8 Step SStep
      LabR(j).Visible = True
      PR(j).Visible = True
   Next j
End Sub


Private Sub PR_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   YS = Y
   XS = X
End Sub

Private Sub PR_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Index is button 0,1,2,,8
' ReDim LX(0 To 2, 0 To 8)
' ReDim LY(0 To 2, 0 To 8)
' ReDim PalRGB(0 To 2, 0 To 255) As Long
' PalRGB(0, LY(cul,0)-LY(cul,1)-LY(cul,2)) linear
' SCUL = Selected color R,G or B

Dim iy As Long
Dim SStep As Long
   If Button = vbLeftButton Then
      iy = PR(Index).Top + (Y - YS)
      If iy < 0 Then iy = 0
      If iy > 255 Then iy = 255
      PR(Index).Top = iy
      LY(SCUL, Index) = iy
      LabR(Index) = 255 - iy
      SStep = 8 \ (NumButtons(SCUL) - 1)  '2,3,5,9 -> 8,4,2,1
      AvoidButtonCrossing SStep, Index, X, XS
      DrawLines SStep
      GetRGBS SStep
      ShowPAL
   End If
End Sub

Private Sub AvoidButtonCrossing(SStep As Long, TheIndex As Integer, X As Single, XS As Single)
Dim k As Long
   For k = SStep To 8 - SStep Step SStep
      If k = TheIndex Then
         LX(SCUL, k) = PR(k).Left + (X - XS)
         If LX(SCUL, k) < LX(SCUL, k - SStep) Then LX(SCUL, k) = LX(SCUL, k - SStep)
         If LX(SCUL, k) > LX(SCUL, k + SStep) Then LX(SCUL, k) = LX(SCUL, k + SStep)
         PR(k).Left = LX(SCUL, k)
      End If
   Next k
End Sub

Private Sub DrawLines(SStep As Long)
Dim k As Long
   PIC.Cls
   For k = 0 To 8 - SStep Step SStep
      PIC.Line (LX(SCUL, k), LY(SCUL, k))-(LX(SCUL, k + SStep), LY(SCUL, k + SStep)), ButtonCul(SCUL)
   Next k
End Sub

Private Sub GetRGBS(SStep As Long)
' ReDim PalRGB(0 To 2, 0 To 255) As Long
' ReDim LX(0 To 2, 0 To 8)
' ReDim LY(0 To 2, 0 To 8)
Dim m As Single, c As Single
Dim dx As Single, dy As Single
Dim k As Long, j As Long
   ' y = m.x + c
   ' c = y - m.x
   ' IE Linear interpolation
   ' EG for 3 buttons
   ' x0=LX(cul,0),   y0=LY(Cul,0)
   ' x1=LX(Cul,1),   y1=LY(Cul,1)
   ' x2=LX(Cul,2),   y2=LY(Cul,2)
   ' m = (y1-y0)/(x1-x0):  c = y0 - m.xo
   ' m = (y2-y1)/(x2-x1):  c = y1 - m.x1
   ' PalRGB(0, LY(cul,0)-LY(cul,1)-LY(cul,2)) linear
   For j = 0 To 8 - SStep Step SStep ' EG 0-4
      dx = (LX(SCUL, j + SStep) - LX(SCUL, j))
      dy = (LY(SCUL, j + SStep) - LY(SCUL, j))
      If dx > 0 Then
         m = dy / dx
      Else
         m = 10000 * Sgn(dy)
      End If
      c = LY(SCUL, j) - m * LX(SCUL, j)
      For k = LX(SCUL, j) To LX(SCUL, j + SStep) ' EG Button0 -> Button4 xpos
         PalRGB(SCUL, k) = m * k + c
         If PalRGB(SCUL, k) < 0 Then PalRGB(SCUL, k) = 0
         If PalRGB(SCUL, k) > 255 Then PalRGB(SCUL, k) = 255
         PalRGB(SCUL, k) = 255 - PalRGB(SCUL, k) ' Invert
      Next k
   Next j
End Sub

Private Sub ShowPAL()
Dim k As Long
   For k = 0 To 255  ' Show palette from PalRGB()
      PICPAL.Line (k, 0)-(k, PICPAL.Height), RGB(PalRGB(0, k), PalRGB(1, k), PalRGB(2, k))
   Next k
End Sub
         
Private Function FileExists(FSpec$) As Boolean
  If Dir$(FSpec$) <> "" Then FileExists = True
End Function

Private Function GetFileName(FSpec$) As String
Dim L As Long
Dim k As Long
Dim A$
   GetFileName = ""
   L = Len(FSpec$)
   If L < 1 Then Exit Function
   For k = L To 1 Step -1
      A$ = Mid$(FSpec$, k, 1)
      If Mid$(FSpec$, k, 1) = "\" Then Exit For
   Next k
   If k = 0 Then
      GetFileName = FSpec$
   Else
      GetFileName = Right$(FSpec$, L - k)
   End If
End Function
         
Private Sub cmdFileOps_Click(Index As Integer)
   Select Case Index
   Case 0   ' Load BPS file
      LoadBPS
   Case 1   ' Save BPS file
      SaveBPS
   Case 2   ' Open PAL file
      LoadPAL
   Case 3   ' Save JASC PAL
      SavePAL
   End Select
End Sub

Private Sub LoadBPS()
Dim k As Long, j As Long
Dim p As Long
Dim A$
   On Error GoTo LoadError
   With CD
      .DialogTitle = "Open BPS file"
      .DefaultExt = ".bps"
      .InitDir = FileSpec$
      .FileName = ""
      .Filter = "Button positions(*.bps)|*.bps"
      .ShowOpen
      FileSpec$ = .FileName
   End With
   
   If FileSpec$ <> "" Then
      ' ReDim NumButtons(0 To 2)   ' #R,#G,#B
      ' ReDim LX(0 To 2, 0 To 8)   ' Cul 0,1,2, xpos button 0,1,2
      ' ReDim LY(0 To 2, 0 To 8)   ' Cul 0,1,2, ypos button 0,1,2
      ' Save for Error
      pNumButtons() = NumButtons()
      pLX() = LX()
      pLY() = LY()
      
      p = FreeFile
      Open FileSpec$ For Input As #p
      For k = 0 To 2
         Input #p, NumButtons(k)
      Next k
      For j = 0 To 2
         For k = 0 To 8
            Input #p, LX(j, k)
         Next k
      Next j
      For j = 0 To 2
         For k = 0 To 8
            Input #p, LY(j, k)
         Next k
      Next j
      Close #p
      '
      ' LoadSetter
      For j = 2 To 0 Step -1
         k = 1 + Log(NumButtons(j) - 1) / Log(2) ' ' N=2,3,5,9 -> k=1,2,3,4
         SCUL = j   ' Selected color
         HSNButtons(CInt(j)).Value = k
         Call HSNButtons_Change(CInt(j))
         RGBChanger
      Next j
      optCul(0) = True
   
   End If
   picTEST.BackColor = 0
   On Error GoTo 0
   A$ = GetFileName(FileSpec$)
   Caption = "  Palette maker:  " & LCase$(A$)
   Exit Sub
'=============
LoadError:
   Close #p
   MsgBox FileSpec$ & " ERROR", vbCritical, "Opening BPS file"
   
   NumButtons() = pNumButtons()
   LX() = pLX()
   LY() = pLY()
   ' Prev LoadSetter
   For j = 2 To 0 Step -1
      k = 1 + Log(NumButtons(j) - 1) / Log(2) ' ' N=2,3,5,9 -> k=1,2,3,4
      SCUL = j   ' Selected color
      HSNButtons(CInt(j)).Value = k
      Call HSNButtons_Change(CInt(j))
      RGBChanger
   Next j
   optCul(0) = True
   picTEST.BackColor = 0
   Caption = "  Palette maker"
End Sub

Private Sub LoadPAL()
Dim k As Long, j As Long
Dim p As Long
Dim A$
Dim s As Long
Dim Sum As Long
   On Error GoTo LoadPALError
   With CD
      .DialogTitle = "Open PAL file"
      .DefaultExt = ".pal"
      .InitDir = FileSpec$
      .FileName = ""
      .Filter = "Palette(*.pal)|*.pal"
      .ShowOpen
      FileSpec$ = .FileName
   End With
   
   If FileSpec$ <> "" Then
      ' Save for Error
      pNumButtons() = NumButtons()
      pLX() = LX()
      pLY() = LY()
      
      p = FreeFile
      Open FileSpec$ For Input As #p
      Line Input #p, A$
      If A$ <> "JASC-PAL" Then GoTo LoadPALError
      Line Input #p, A$ ' "0100"
      Line Input #p, A$
      If Val(A$) <> 256 Then
         MsgBox "Not a 256 PAL file", vbCritical, "Opening PAL file"
         GoTo LoadPALError
      End If
      ' Now 256 PAL
      ReDim PalRGB(0 To 2, 0 To 255)   ' Palette - 0, Black
      For k = 0 To 255
         Input #p, PalRGB(0, k), PalRGB(1, k), PalRGB(2, k)
      Next k
      Close #p
      
      For k = 0 To 255  ' Show palette from PalRGB()
         picTEST.Line (k, 0)-(k, picTEST.Height), RGB(PalRGB(0, k), PalRGB(1, k), PalRGB(2, k))
      Next k
      
      ' Assign average values to 9 RGB buttons
      ' eg PalRGB(0,0-255)   ' Red
      For j = 0 To 2
         Sum = 0
         For k = 0 To 15
            Sum = Sum + PalRGB(j, k)
         Next k
         LY(j, 0) = 255 - (Sum \ 16)
      Next j
      For j = 0 To 2
         For p = 1 To 7
            s = 16 * (2 * p - 1)
            Sum = 0
            For k = s To s + 31
               Sum = Sum + PalRGB(j, k)
            Next k
            LY(j, p) = 255 - (Sum \ 32)
         Next p
      Next j
      For j = 0 To 2
         Sum = 0
         For k = 240 To 255
            Sum = Sum + PalRGB(j, k)
         Next k
         LY(j, 8) = 255 - (Sum \ 16)
      Next j
      ' Set 9 LX()s
      For j = 0 To 8
         For k = 0 To 2
            If j = 8 Then
               LX(k, j) = 255
            Else
               LX(k, j) = j * 32
            End If
         Next k
      Next j
      
      ' LoadSetter
      For j = 2 To 0 Step -1
         SCUL = j   ' Selected color
         HSNButtons(CInt(j)).Value = 4  ' Make 9 buttons
         Call HSNButtons_Change(CInt(j))
         RGBChanger
      Next j
      optCul(0) = True
      'ShowPAL ' Test
   End If
   On Error GoTo 0
   A$ = GetFileName(FileSpec$)
   Caption = "  Palette maker:  " & LCase$(A$)
   Exit Sub
'=============
LoadPALError:
   Close #p
   MsgBox FileSpec$ & " ERROR", vbCritical, "Opening PAL file"
   NumButtons() = pNumButtons()
   LX() = pLX()
   LY() = pLY()
   ' Prev LoadSetter
   For j = 2 To 0 Step -1
      k = 1 + Log(NumButtons(j) - 1) / Log(2) ' ' N=2,3,5,9 -> k=1,2,3,4
      SCUL = j   ' Selected color
      HSNButtons(CInt(j)).Value = k
      Call HSNButtons_Change(CInt(j))
      RGBChanger
   Next j
   optCul(0) = True
   Caption = "  Palette maker"
End Sub

Private Sub SavePAL()
Dim Ext$
Dim p As Long, k As Long
Dim R$, G$, B$
   With CD
      .DialogTitle = "Save JASC PAL file"
      .DefaultExt = ".pal"
      .InitDir = FileSpec$
      .FileName = ""
      .Flags = &H2   ' Checks if file exists
      .Filter = "Palette(*.pal)|*.pal"
      .ShowSave
      FileSpec$ = .FileName
   End With
   
   ' If FileSpec$ has no ext  (ie no .) then .pal added
   
   If FileSpec$ <> "" Then
      ' Check extension
      p = InStr(1, FileSpec$, ".")
      Ext$ = LCase$(Mid$(FileSpec$, p))
      If Ext$ <> ".pal" Then
         p = MsgBox("File extension = " & Ext$ & vbCrLf & "Continue ?", vbQuestion Or vbYesNo, "File extension")
         If p = vbNo Then
            Caption = "  Palette maker"
            Exit Sub
         End If
      End If
      R$ = GetFileName(FileSpec$)
      Caption = "  Palette maker:  " & LCase$(R$)
      ' JASC-PAL
      ' 0100
      ' 256
      ' 0 0 0
      ' 0 0 0
      p = FreeFile
      Open FileSpec$ For Output As #p
      Print #p, "JASC-PAL"
      Print #p, "0100"
      Print #p, "256"
      ' PalRGB(0, k), PalRGB(1, k), PalRGB(2, k))
      For k = 0 To 255
         R$ = Trim$(Str$(PalRGB(0, k)))
         G$ = Trim$(Str$(PalRGB(1, k)))
         B$ = Trim$(Str$(PalRGB(2, k)))
         Print #p, R$ & " " & G$ & " " & B$
      Next k
      Close #p
   End If
End Sub

Private Sub SaveBPS()
Dim Ext$
Dim k As Long, j As Long
Dim p As Long
   With CD
      .DialogTitle = "Save BPS file"
      .DefaultExt = ".bps"
      .InitDir = FileSpec$
      .FileName = ""
      .Flags = &H2   ' Checks if file exists
      .Filter = "Button positions(*.bps)|*.bps"
      .ShowSave
      FileSpec$ = .FileName
   End With
   
   ' If FileSpec$ has no ext  (ie no .) then .bps added
   
   If FileSpec$ <> "" Then
      ' ReDim NumButtons(0 To 2)   ' #R,#G,#B
      ' ReDim LX(0 To 2, 0 To 8)   ' Cul 0,1,2, xpos button 0,1,2
      ' ReDim LY(0 To 2, 0 To 8)   ' Cul 0,1,2, ypos button 0,1,2
      ' Check extension
      p = InStr(1, FileSpec$, ".")
      Ext$ = LCase$(Mid$(FileSpec$, p))
      If Ext$ <> ".bps" Then
         p = MsgBox("File extension = " & Ext$ & vbCrLf & "Continue ?", vbQuestion Or vbYesNo, "File extension")
         If p = vbNo Then
            Caption = "  Palette maker"
            Exit Sub
         End If
      End If
      p = FreeFile
      Open FileSpec$ For Output As #p
      ' Print NumButtons(0 To 2)   ' #R,#G,#B
      For k = 0 To 2
         If k < 2 Then
            Print #p, NumButtons(k) & ",";
         Else
            Print #p, Trim$(Str$(NumButtons(k)))
         End If
      Next k
      ' Print LX(0 To 2, 0 To 8)
      For j = 0 To 2
      For k = 0 To 8
         If k < 8 Then
            Print #p, LX(j, k) & ",";
         Else
            Print #p, Trim$(Str$(LX(j, k)))
         End If
      Next k
      Next j
      
      ' Print LY(0 To 2, 0 To 8)
      ' Trebor Tnemyar
      For j = 0 To 2
      For k = 0 To 8
         If k < 8 Then
            Print #p, LY(j, k) & ",";
         Else
            Print #p, Trim$(Str$(LY(j, k)))
         End If
      Next k
      Next j
      Close #p
   End If
End Sub


