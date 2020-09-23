VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7500
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   165
      Top             =   5445
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "CD"
   End
   Begin VB.CommandButton cmdPalMaker 
      Caption         =   "Palette maker"
      Height          =   390
      Left            =   2340
      TabIndex        =   22
      Top             =   4905
      Width           =   1500
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Image"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   4905
      Width           =   1395
   End
   Begin VB.Frame Frame3 
      Height          =   5730
      Left            =   4185
      TabIndex        =   2
      Top             =   75
      Width           =   3105
      Begin VB.CheckBox chkAnim 
         Caption         =   "Cycle"
         Height          =   225
         Index           =   3
         Left            =   315
         TabIndex        =   33
         Top             =   5040
         Width           =   1065
      End
      Begin VB.HScrollBar HSSpeeds 
         Height          =   195
         Index           =   3
         Left            =   1545
         Max             =   10
         Min             =   -10
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   5130
         Value           =   1
         Width           =   570
      End
      Begin VB.CommandButton cmdGO 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Sphere + Ring"
         Height          =   330
         Index           =   2
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3270
         Width           =   1410
      End
      Begin VB.HScrollBar HSSpeeds 
         Height          =   195
         Index           =   2
         Left            =   1545
         Max             =   10
         Min             =   -10
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4830
         Value           =   1
         Width           =   570
      End
      Begin VB.HScrollBar HSSpeeds 
         Height          =   195
         Index           =   1
         Left            =   1545
         Max             =   10
         Min             =   -10
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   4545
         Value           =   1
         Width           =   570
      End
      Begin VB.HScrollBar HSSpeeds 
         Height          =   195
         Index           =   0
         Left            =   1545
         Max             =   10
         Min             =   -10
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4260
         Value           =   1
         Width           =   570
      End
      Begin VB.CheckBox chkAnim 
         Caption         =   "Rotate Y"
         Height          =   225
         Index           =   2
         Left            =   315
         TabIndex        =   21
         Top             =   4740
         Width           =   1065
      End
      Begin VB.CommandButton cmdGO 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Tunnel"
         Height          =   330
         Index           =   3
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3660
         Width           =   1410
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   270
         Left            =   285
         TabIndex        =   19
         Top             =   5355
         Width           =   990
      End
      Begin VB.CheckBox chkAnim 
         Caption         =   "Rotate Z"
         Height          =   225
         Index           =   1
         Left            =   315
         TabIndex        =   18
         Top             =   4470
         Width           =   1080
      End
      Begin VB.CheckBox chkAnim 
         Caption         =   "Scroll"
         Height          =   225
         Index           =   0
         Left            =   315
         TabIndex        =   16
         Top             =   4200
         Width           =   885
      End
      Begin VB.CommandButton cmdGO 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Sphere"
         Height          =   330
         Index           =   1
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2880
         Width           =   1410
      End
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   375
         Width           =   1290
      End
      Begin VB.CommandButton cmdGO 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cylinder"
         Height          =   330
         Index           =   0
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2490
         Width           =   1410
      End
      Begin VB.Frame Frame2 
         Height          =   2205
         Left            =   120
         TabIndex        =   3
         Top             =   165
         Width           =   1425
         Begin VB.HScrollBar HScroll1 
            Height          =   210
            Index           =   0
            LargeChange     =   8
            Left            =   180
            Max             =   128
            Min             =   1
            TabIndex        =   7
            Top             =   465
            Value           =   1
            Width           =   1110
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   210
            Index           =   1
            LargeChange     =   16
            Left            =   180
            Max             =   128
            Min             =   4
            SmallChange     =   4
            TabIndex        =   6
            Top             =   1020
            Value           =   4
            Width           =   1110
         End
         Begin VB.CheckBox chkWrap 
            Caption         =   "Wrap X"
            Height          =   240
            Index           =   0
            Left            =   300
            TabIndex        =   5
            Top             =   1410
            Width           =   960
         End
         Begin VB.CheckBox chkWrap 
            Caption         =   "Wrap Y"
            Height          =   240
            Index           =   1
            Left            =   300
            TabIndex        =   4
            Top             =   1725
            Width           =   990
         End
         Begin VB.Label LabNum 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   990
            TabIndex        =   11
            Top             =   210
            Width           =   315
         End
         Begin VB.Label Label1 
            Caption         =   "Graininess = "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   10
            Top             =   225
            Width           =   810
         End
         Begin VB.Label LabNum 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   1
            Left            =   900
            TabIndex        =   9
            Top             =   735
            Width           =   315
         End
         Begin VB.Label Label1 
            Caption         =   "Scale = "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   8
            Top             =   750
            Width           =   675
         End
      End
      Begin VB.Label LabSpeeds 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   2265
         TabIndex        =   32
         Top             =   5085
         Width           =   315
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   2670
         Picture         =   "Main.frx":0442
         Top             =   5295
         Width           =   360
      End
      Begin VB.Label LabSpeeds 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   2265
         TabIndex        =   29
         Top             =   4800
         Width           =   315
      End
      Begin VB.Label LabSpeeds 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   2265
         TabIndex        =   28
         Top             =   4515
         Width           =   315
      End
      Begin VB.Label LabSpeeds 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   2265
         TabIndex        =   27
         Top             =   4230
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "Speeds"
         Height          =   225
         Left            =   1830
         TabIndex        =   24
         Top             =   3990
         Width           =   750
      End
      Begin VB.Label LabPalName 
         Caption         =   "PalName"
         Height          =   210
         Left            =   1680
         TabIndex        =   14
         Top             =   150
         Width           =   1155
      End
   End
   Begin VB.PictureBox picPal 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   135
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   1
      Top             =   4365
      Width           =   3900
   End
   Begin VB.PictureBox PIC 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   180
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   240
      Width           =   3840
   End
   Begin VB.Shape Shape1 
      Height          =   4005
      Left            =   120
      Top             =   180
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Main.frm

' Plasma Play:  by Robert Rayment June 2006

Option Explicit

Option Base 1

'--- This lot to blit array to picbox -----------------------------------------------------------------------
Private Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Private Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Private Type BITMAPINFO
   bmiHeader As BITMAPINFOHEADER
   bmiColors As RGBQUAD
End Type
Dim BB As BITMAPINFO

Private Declare Function SetDIBits Lib "gdi32.dll" _
(ByVal hdc As Long, _
ByVal hBitmap As Long, _
ByVal nStartScan As Long, _
ByVal nNumScans As Long, _
ByRef lpBits As Any, _
ByRef lpBI As BITMAPINFO, _
ByVal wUsage As Long) As Long

Private Const DIB_PAL_COLORS As Long = 1
Private Const DIB_RGB_COLORS As Long = 0
'---------------------------------------------------------------------------


'---- This used for scrolling array ----------------------------------------
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
(ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'---------------------------------------------------------------------------

'---- This used in some Do Loops to reduce CPU usgae -----------------------
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
'---------------------------------------------------------------------------


Private sizex As Integer, sizey As Integer ' PIC size

' Plasma parameters
Private StartNoise
Private StartStepsize
Private WrapX As Boolean
Private WrapY As Boolean

' RGB components
Private red As Byte
Private green As Byte
Private blue As Byte

Private PaletteMaxIndex           ' Max Palette size

' Public
' PalRGB() As Long
' IndexArray() As Integer   ' To hold palette indexes
' ColArray() As Long        ' To hold colors for displaying
' Colors()                  ' The 512 palette colors

Private stColArray() As Long
Private BackArray() As Long

' For tunnel
Private CoordsX() As Long
Private CoordsY() As Long

Private Effects As Integer
Private aScroll As Boolean
Private aRotate As Boolean
Private aRotateVert As Boolean
Private Actions As Long
Private aChkChange As Boolean
Private aTunnelDone As Boolean
Private aCycle As Boolean
Private aPlasmaDone As Boolean


Private ScrollSpeed As Long
Private RotSpeed As Long
Private RotVSpeed As Long
Private CycleSpeed As Long

' LUTs for vertical rotations
' PRECALC
Private xindent() As Single
Private LookVH() As Long
Private LookSX() As Long
Private LookSY() As Long
Private LookPX() As Long
Private LookPY() As Long


Private PathSpec$, PalDir$, FileSpec$
Const pi# = 3.14159265
Const dtr# = pi# / 180



Private Sub cmdPalMaker_Click()
   cmdStop_Click
   frmPaletteMaker.Show
End Sub


Private Sub Form_Initialize()
   '++ For manifest file +++++++++++++
   m_hMod = LoadLibrary("shell32.dll")
   InitCommonControls
   '++++++++++++++++++++++++++++++++++
   ' PIC size
   sizex = 256
   sizey = 256
   ReDim IndexArray(1 To sizex, 1 To sizey)
   ReDim ColArray(1 To sizex, 1 To sizey)
   ReDim stColArray(1 To sizex, 1 To sizey)
   PaletteMaxIndex = 511
   ReDim Colors(0 To PaletteMaxIndex)
   ReDim BackArray(1 To sizex, 1 To sizey)
   PRECALC
End Sub

Private Sub Form_Load()
Dim A$, Spec$, PalSpec$
   
'   If (App.LogMode <> 1) Then
'    MsgBox "Faster if compiled", vbExclamation, "Plasma Play"
'   End If
   
   Caption = " Plasma Play by Robert Rayment"
   
   Show
   
   ' Locations
   With PIC
      .Width = sizex
      .Height = sizey
      .Top = 12
      .Left = 12
   End With
   
   ' PIC border
   With Shape1
      .Top = PIC.Top - 2
      .Left = PIC.Left - 2
      .Width = sizex + 4
      .Height = sizey + 4
   End With
   
   ' Pic to show palette
   picPal.Top = PIC.Top + PIC.Height + 12
   picPal.Left = PIC.Left - 2
   picPal.Width = sizex + 4
   
   cmdSave.Top = picPal.Top + picPal.Height + 10
   cmdPalMaker.Top = picPal.Top + picPal.Height + 10
   
   ' Initial plasma parameters
   
   StartNoise = 0
   StartStepsize = 32
   
   HScroll1(0).Value = 1    ' Graininess
   LabNum(0).Caption = "1"
   HScroll1(1).Value = 32  ' Scale
   LabNum(1).Caption = "32"
   
   chkWrap(0).Value = Checked
   chkWrap(1).Value = Checked
   WrapX = True
   WrapY = True
   
   cmdGO(0).BackColor = RGB(192, 255, 192)   ' Light Green
   cmdGO(1).BackColor = RGB(192, 255, 192)   ' Light Green
   
   
   'Get app path
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' GET PAL FILES
   ' If Pals folder not in with App Folder
   ' or compiled exe not in App Folder will
   ' need to alter this code here
   
   ' To pick up *.pal files from Pal folder in app folder
   '------------------
   'Get FIRST entry
   PalDir$ = PathSpec$ & "\Pals_bps\"
   Spec$ = PalDir$ & "*.pal"
   A$ = Dir$(Spec$)
   '------------------
   
   If A$ = "" Then
      MsgBox "Pal files not found"
      Exit Sub
   End If
   
   List1.AddItem UCase$(Left$(A$, 1)) & LCase$(Mid$(A$, 2))
   Do
     A$ = Trim$(Dir)    'Gets NEXT entry
     If A$ <> "" Then
        List1.AddItem UCase$(Left$(A$, 1)) & LCase$(Mid$(A$, 2))
     Else   'A$="" indicates no more entries
        Exit Do
     End If
   Loop
   
   ' Starting palette
   List1.Selected(0) = True
   LabPalName = List1.List(0)
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   PalSpec$ = PalDir$ & LabPalName.Caption
   ReadPAL PalSpec$
   
   Load frmPaletteMaker
   frmPaletteMaker.Hide
   
   ' Prepare the bitmap description.
   With BB.bmiHeader
         .biSize = 40
         .biWidth = sizex
         .biHeight = -sizey
         .biPlanes = 1
         .biBitCount = 32
   End With
   
   Actions = 0
   Effects = -1
   
   HSSpeeds(0).Value = 1
   HSSpeeds(1).Value = 1
   HSSpeeds(2).Value = 1
   HSSpeeds(3).Value = 1
   
   ScrollSpeed = 1
   RotSpeed = 1
   RotVSpeed = 1
   CycleSpeed = 1
End Sub

Private Sub HSSpeeds_Change(Index As Integer)
   Select Case Index
   Case 0   ' ScrollSpped
      ScrollSpeed = HSSpeeds(0).Value
      LabSpeeds(0) = ScrollSpeed
   Case 1   ' RotSpeed
      RotSpeed = HSSpeeds(1).Value
      LabSpeeds(1) = RotSpeed
   Case 2   ' RotVSpeed
      RotVSpeed = HSSpeeds(2).Value
      LabSpeeds(2) = RotVSpeed
   Case 3   ' CycleSpeed
      CycleSpeed = HSSpeeds(3).Value
      LabSpeeds(3) = CycleSpeed
   End Select
End Sub

Private Sub cmdStop_Click()
   chkAnim(0).Value = Unchecked
   aScroll = False
   chkAnim(1).Value = Unchecked
   aRotate = False
   chkAnim(2).Value = Unchecked
   aRotateVert = False
   chkAnim(3).Value = Unchecked
   aCycle = False
   aTunnelDone = True
End Sub


'###### GO GO GO #############################################

Private Sub chkAnim_Click(Index As Integer)
   aChkChange = True
   Select Case Index
   Case 0   ' Scroll
      aScroll = -chkAnim(0).Value
      aTunnelDone = True
   Case 1   ' Rotate in plane
      aRotate = -chkAnim(1).Value
      aTunnelDone = True
   Case 2   ' Rotate about vertical
      aRotateVert = -chkAnim(2).Value
      aTunnelDone = True
   Case 3   ' Cycle
      aCycle = -chkAnim(3).Value
      aTunnelDone = True
   End Select

'   Actions = Abs(chkAnim(0).Value) + 2 * (Abs(chkAnim(1).Value)) + 4 * (Abs(chkAnim(2).Value))
'   cmdGO_Click Effects
End Sub


Private Sub cmdGO_Click(Index As Integer)
Dim k As Long
   Effects = Index
   Actions = Abs(chkAnim(0).Value) + 2 * (Abs(chkAnim(1).Value)) + 4 * (Abs(chkAnim(2).Value))
   Actions = Actions + 8 * Effects
   aChkChange = False

'   Select Case Actions
'   Cylinder,Sphere,Sphere+Ring,Tunnel
'   Case 0, 8,16,24   ' None
'   Case 1, 9,17,25   ' Scroll
'   Case 2,10,18,26   ' Rotate
'   Case 3,11,19,27   ' Scroll & Rotate
'   Case 4,12,20,28   ' RotateVert
'   Case 5,13,21,29   ' RotateVert & Scroll
'   Case 6,14,22,30   ' RotateVert & Rotate
'   Case 7,15,23,31   ' RotateVert & Scroll & Rotate  X
'   End Select
   
   For k = 0 To 3
      If k = Index Then
         cmdGO(k).BackColor = RGB(255, 192, 192)
      Else
         cmdGO(k).BackColor = RGB(192, 255, 192)
      End If
      If Index <> 3 Then aTunnelDone = True
   Next k
   
   DisableControls
   
   If aCycle Then
      Call CycleAction
      EnableControls
      cmdGO(Index).BackColor = RGB(192, 255, 192)
      Exit Sub
   End If
   
   Plasma
   
   Select Case Actions
   Case 0 To 7
      Call CyinderActions
   Case 8 To 23
      Call SphereActions   ' >= 16 SphereRingActions
'   Case 16 To 23
'      Call SphereRingActions
   Case 24 To 31
      Call TunnelActions
   End Select
   
   EnableControls
   
   cmdGO(Index).BackColor = RGB(192, 255, 192)
End Sub


'######## ACTIONS ##############################################

Private Sub CycleAction()
Dim ix As Long, iy As Long
Dim ixsc As Long
Dim LCul As Long
Dim R As Long, G As Long, B As Long
      
   chkAnim(0).Value = Unchecked
   aScroll = False
   chkAnim(1).Value = Unchecked
   aRotate = False
   chkAnim(2).Value = Unchecked
   aRotateVert = False
   aTunnelDone = True
   If Not aPlasmaDone Then Plasma
   Select Case Actions
   Case 0 To 7
     BackArray() = ColArray()
     DisplayBackArray
   Case 8 To 15
      ixsc = sizex \ 2
      GetSphere ixsc
      DisplayBackArray
   Case 16 To 23
      ixsc = sizex \ 2
      GetSphere ixsc
      DisplayBackArray
   Case 24 To 31
      TunnelActions
   End Select
   
   Do
      For iy = sizey To 1 Step -1
      For ix = sizex To 1 Step -1
         LCul = BackArray(ix, iy)
         If LCul <> 0 Then
            R = LCul And &HFF&
            G = (LCul And &HFF00&) \ &H100&
            B = (LCul And &HFF0000) \ &H10000
            R = (R + CycleSpeed) ' And 255
            G = (G + CycleSpeed) ' And 255
            B = (B + CycleSpeed) ' And 255
            If R > 255 Then R = 1
            If G > 255 Then G = 1
            If B > 255 Then B = 1
            If R < 0 Then R = 255
            If G < 0 Then G = 255
            If B < 0 Then B = 255
            BackArray(ix, iy) = RGB(R, G, B)
         End If
         If Not aCycle Then Exit For
      Next ix
         If Not aCycle Then Exit For
      Next iy
      DisplayBackArray
      DoEvents
      Sleep 1
   Loop Until Not aCycle
End Sub

Private Sub CyinderActions()
Dim ixsc As Long
Dim zang As Single
Dim k As Long
   Select Case Actions
   Case 0   ' CYLINDER None
        BackArray() = ColArray()
        DisplayBackArray
   Case 1   ' CYLINDER Scroll
         Do
            Scroll
            DisplayBackArray
            DoEvents
            If aChkChange Then Exit Do
            Sleep 1
         Loop Until Not aScroll
   Case 2   ' CYLINDER Rotate
        zang = 0
        Do
            zang = zang + RotSpeed
            If zang > 360 Then zang = zang - 360
            Rotate zang
            DisplayBackArray
            DoEvents
            If aChkChange Then Exit Do
            Sleep 1
         Loop Until Not aRotate
   Case 3   ' CYLINDER Scroll & Rotate
        zang = 0
        Do
            Scroll
            zang = zang + RotSpeed
            If zang > 360 Then zang = zang - 360
            Rotate zang
            DisplayBackArray
            DoEvents
            If aChkChange Then Exit Do
            Sleep 1
         Loop Until Not aRotate
   Case 4   ' CYLINDER RotateVert
         ixsc = 1
         Do
            GetCylinder ixsc
            DisplayBackArray
            ixsc = ixsc + RotVSpeed
            If ixsc > sizex Then ixsc = ixsc - sizex
            If ixsc < 1 Then
               ixsc = sizex + ixsc
            End If
            DoEvents
            If aChkChange Then Exit Do
            Sleep 1
         Loop Until Not aRotateVert
   Case 5   ' CYLINDER RotateVert & Scroll
        ixsc = sizex \ 2
        Do
            Scroll
            GetCylinder ixsc
            DisplayBackArray
            ixsc = ixsc + RotVSpeed
            If ixsc > sizex Then ixsc = ixsc - sizex
            If ixsc < 1 Then
               ixsc = sizex + ixsc
            End If
            DoEvents
            If aChkChange Then Exit Do
            Sleep 1
         Loop Until Not aRotateVert
   Case 6   '  CYLINDER RotateVert & Rotate
        zang = 0
        ixsc = sizex \ 2
        Do
            GetCylinder ixsc   ' ColArray()->BackArray
            ' save ColArray()
            stColArray() = ColArray()
            ColArray() = BackArray()
            zang = zang + RotSpeed
            If zang > 360 Then zang = zang - 360
            Rotate zang        ' ReDim BackArray(1 To sizex, 1 To sizey)
                               ' &  ColArray()->BackArray
            ixsc = ixsc + RotVSpeed
            If ixsc > sizex Then ixsc = ixsc - sizex
            If ixsc < 1 Then
               ixsc = sizex + ixsc
            End If
            DisplayBackArray
            DoEvents
            ' restore ColArray()
            ColArray() = stColArray()
            If aChkChange Then Exit Do
            Sleep 1
         Loop Until Not aRotateVert
   Case 7   ' CYLINDER RotateVert & Scroll & Rotate  X
        BackArray() = ColArray()
        DisplayBackArray
   End Select
End Sub

Private Sub SphereActions()
Dim ixsc As Long
Dim zang As Single
Dim k As Long
   Select Case Actions
   Case 8, 16  ' SPHERE & SPHERE RING None
      ixsc = sizex \ 2
      GetSphere ixsc
      DisplayBackArray
   Case 9, 17  ' SPHERE & SPHERE RING Scroll
      ixsc = sizex \ 2
      Do
         Scroll
         GetSphere ixsc
         DisplayBackArray
         DoEvents
         If aChkChange Then Exit Do
         Sleep 1
      Loop Until Not aScroll
   Case 10, 18  ' SPHERE & SPHERE RING Rotate
        ixsc = sizex \ 2
        GetSphere ixsc
        ColArray() = BackArray()
        zang = 0
        Do
            zang = zang + RotSpeed
            If zang > 360 Then zang = zang - 360
            Rotate zang
            DisplayBackArray
            DoEvents
            If aChkChange Then Exit Do
            Sleep 1
         Loop Until Not aRotate
   Case 11, 19  ' SPHERE & SPHERE RING Scroll & Rotate
      zang = 0
      Do
          Scroll
          ' save ColArray() for next scroll
          stColArray() = ColArray()
          GetSphere ixsc
          ColArray() = BackArray()
          zang = zang + RotSpeed
          If zang > 360 Then zang = zang - 360
          Rotate zang
          DisplayBackArray
          ' restore ColArray()
          ColArray() = stColArray()
          DoEvents
          If aChkChange Then Exit Do
          Sleep 1
      Loop Until Not aRotate Or Not aScroll
   Case 12, 20  ' SPHERE & SPHERE RING RotateVert
      ixsc = sizex \ 2
      Do
         GetSphere ixsc
         DisplayBackArray
         ixsc = ixsc + RotVSpeed
         If ixsc > sizex Then ixsc = ixsc - sizex
         If ixsc < 1 Then ixsc = sizex + ixsc
         DoEvents
         If aChkChange Then Exit Do
         Sleep 1
      Loop Until Not aRotateVert
   Case 13, 21  ' SPHERE & SPHERE RING RotateVert & Scroll
      ixsc = sizex \ 2
      Do
         Scroll
         GetSphere ixsc
         DisplayBackArray
         ixsc = ixsc + RotVSpeed
         If ixsc > sizex Then ixsc = ixsc - sizex
         If ixsc < 1 Then ixsc = sizex + ixsc
         DoEvents
         If aChkChange Then Exit Do
         Sleep 1
      Loop Until Not aRotateVert Or Not aScroll
   Case 14, 22  ' SPHERE & SPHERE RING RotateVert & Rotate
        zang = 0
        ixsc = sizex \ 2
        Do
            GetSphere ixsc   ' ColArray()->BackArray
            ' save ColArray()
            stColArray() = ColArray()
            ColArray() = BackArray()
            zang = zang + RotSpeed
            If zang > 360 Then zang = zang - 360
            Rotate zang        ' ReDim BackArray(1 To sizex, 1 To sizey)
                               ' &  ColArray()->BackArray
            ixsc = ixsc + RotVSpeed
            If ixsc > sizex Then ixsc = ixsc - sizex
            If ixsc < 1 Then
               ixsc = sizex + ixsc
            End If
            DisplayBackArray
            DoEvents
            ' restore ColArray()
            ColArray() = stColArray()
            If aChkChange Then Exit Do
            Sleep 1
         Loop Until Not aRotateVert
   Case 15, 23  ' SPHERE & SPHERE RING RotateVert & Scroll & Rotate  X
      ixsc = sizex \ 2
      GetSphere ixsc
      DisplayBackArray
   End Select
End Sub

Private Sub TunnelActions()
Dim kx As Long, ky As Long
Dim ixs As Long, iys As Long
Dim kxstart As Long, kystart As Long
Dim LCul As Long
Dim zr As Long
   ' Plasma done
   aTunnelDone = False
   kxstart = 1
   kystart = 1
   Do
      ' LUT Projection
      For ky = sizey To 1 Step -1
      For kx = sizex To 1 Step -1
         iys = CoordsY(kx, ky) - kystart  ' Using pre-calculated
         ixs = CoordsX(kx, ky) - kxstart  ' CoordsX/Y()
         
         If iys <= 0 Then iys = sizey + iys
         If iys > sizey Then iys = iys - sizey
         If ixs <= 0 Then ixs = sizex + ixs
         If ixs > sizex Then ixs = ixs - sizex
            ' Need to swap R & B
            ' ColArray(kx, ky) = Colors(IndexArray(ixs, iys))
            LCul = Colors(IndexArray(ixs, iys))
            ' R = LCul And &HFF&
            ' G = (LCul And &HFF00&) \ &H100&
            ' B = (LCul And &HFF0000) \ &H10000
            BackArray(kx, ky) = RGB((LCul And &HFF0000) \ &H10000, _
                                    (LCul And &HFF00&) \ &H100&, _
                                     LCul And &HFF&)
'            ' Black center - slow
'            zr = Sqr((ky - sizey \ 2) ^ 2 + (kx - sizex \ 2) ^ 2)
'            If zr < 10 + 4 * Rnd Then
'            'If ky = sizey \ 2 And kx = sizex \ 2 Then
'               ColArray(kx, ky) = 0
'            End If
      Next kx
      Next ky
      
      If aRotateVert And RotVSpeed <> 0 Then ' Advance as X ?
         kxstart = kxstart + RotVSpeed
         If kxstart > sizex Then kxstart = kxstart - sizex
         If kxstart <= 0 Then kxstart = sizex + kxstart
      End If
      
      If aScroll And ScrollSpeed <> 0 Then   ' Advance Y
         kystart = kystart + ScrollSpeed
         If kystart > sizey Then kystart = kystart - sizey
         If kystart <= 0 Then kystart = sizey + kystart
      End If
      
      If aRotate And RotSpeed <> 0 Then      ' Advance X
         kxstart = kxstart + RotSpeed
         If kxstart > sizex Then kxstart = kxstart - sizex
         If kxstart <= 0 Then kxstart = sizex + kxstart
      End If
      
      DisplayBackArray
      DoEvents
      Sleep 1
      If aCycle Then Exit Do
   Loop Until aTunnelDone
End Sub


'###### PLASMA ################################################

Private Sub Plasma()
' ReDim IndexArray(1 To sizex, 1 To sizey)
' ReDim ColArray(1 To sizex, 1 To sizey)
' PaletteMaxIndex = 511
' ReDim Colors(0 To PaletteMaxIndex)
Dim Noise As Long
Dim Stepsize As Long
Dim ix As Long, iy As Long
Dim RndNoise As Single
Dim Col As Long
Dim kx As Long, ky As Long
Dim smax As Long, smin As Long
Dim sdiv As Long
Dim zmul As Single
Dim Index As Integer
   
   ' CORE PLASMA ROUTINE
   
   ReDim IndexArray(1 To sizex, 1 To sizey)
   ReDim ColArray(1 To sizex, 1 To sizey)
   
   Randomize Timer
   
   ' PaletteMaxIndex = 511
   ' Graininess & scale also depends on palette
   ' Rapidly changing palettes will be grainier
   ' StartNoise    = 1,2,,128  ' graininess
   ' StartStepsize = 4,6,,128  ' scale
   
   Noise = StartNoise         ' ie Graininess
   Stepsize = StartStepsize   ' ie Scale
   
   RndNoise = 0
   
   ' Find Col max/min
   smax = -10000
   smin = 10000
   
   '+.... This bit adapted & extended from a QB prog by Alan King 1996
   Do
      
      For iy = 1 To sizey Step Stepsize
      For ix = 1 To sizex Step Stepsize
         
         If Noise > 0 Then RndNoise = (Rnd * (2 * Noise) - Noise) * Stepsize
          
         If iy + Stepsize >= sizey Then
            If WrapY Then
              Col = IndexArray(ix, 1) + RndNoise   ' Gives vertical wrapping
            Else
              Col = IndexArray(ix, sizey) + RndNoise
            End If
         
         ElseIf ix + Stepsize >= sizex Then
            If WrapX Then
               Col = IndexArray(1, iy) + RndNoise  ' Gives horizontal wrapping
            Else
               Col = IndexArray(sizex, iy) + RndNoise
            End If
         
         ElseIf Stepsize = StartStepsize Then
            Col = Rnd * (2 * PaletteMaxIndex) - PaletteMaxIndex
         
         Else
             Col = IndexArray(ix, iy)
             Col = Col + IndexArray(ix + Stepsize, iy)
             Col = Col + IndexArray(ix, iy + Stepsize)
             Col = Col + IndexArray(ix + Stepsize, iy + Stepsize)
             Col = Col / 4 + RndNoise
         End If
         
         If Col > 32767 Then Col = 32767
         If Col < -32768 Then Col = -32768
         
         If Stepsize >= 1 Then 'BoxLine
             
             For ky = iy To iy + Stepsize
                 If ky <= sizey Then
                     For kx = ix To ix + Stepsize
                         If kx <= sizex Then IndexArray(kx, ky) = Col
                     Next kx
                 End If
             Next ky
         
         End If
         If Stepsize < 1 Then IndexArray(ix, iy) = Col
      
         If Col < smin Then smin = Col
         If Col > smax Then smax = Col
      
      Next ix
      Next iy
   
      Stepsize = Stepsize \ 2
   
   Loop Until Stepsize <= 1
   '.........................................................
   
   ' Scale colors indices to range 0 - PaletteMaxIndex
   ' & put full RGB() long colors into ColArray()
   ' & keep new IndexArray()
   
   sdiv = (smax - smin)
   If sdiv <= 0 Then sdiv = 1
   zmul = (PaletteMaxIndex + 1) / sdiv
   
   For iy = sizey To 1 Step -1
   For ix = sizex To 1 Step -1
      
      Index = (IndexArray(ix, iy) - smin) * zmul
      ' Check Index in range
      If Index < 0 Then
         Index = 0
      End If
      If Index > PaletteMaxIndex Then
         Index = PaletteMaxIndex
      End If
      
      ' Put in new Index (nb just changing a palette
      ' will still use the same IndexArray() unless
      ' GO is pressed)
      
      IndexArray(ix, iy) = Index
      ' & fill color array with palette colors
      'ColArray(ix, iy) = Colors(Index)
      ' but reverse R & B SetDIBits used
      Col = Colors(Index)
      'R = LCul And &HFF&
      'G = (LCul And &HFF00&) \ &H100&
      'B = (LCul And &HFF0000) \ &H10000
      ColArray(ix, iy) = RGB((Col And &HFF0000) \ &H10000, _
                              (Col And &HFF00&) \ &H100&, _
                               Col And &HFF&)
   Next ix
   Next iy
   aPlasmaDone = True
End Sub


Private Sub Scroll()
' Move image down
' ReDim ColArray(1 To sizex, 1 To sizey) As Long
' ReDim BackArray(1 To sizex, 1 To sizey)
' Colarray(1 to 256,255)->BackArray(1 to 256,256)
' Colarray(1 to 256,254)->BackArray(1 to 256,255)
'
' Colarray(1 to 256,1)->BackArray(1 to 256,2)
' Colarray(1 to 256,256)->BackArray(1 to 256,1)
Dim iy As Long
Dim SS As Long
   If ScrollSpeed > 0 Then ' Image moves down
      For iy = sizey - ScrollSpeed To 1 Step -1
         CopyMemory BackArray(1, iy + ScrollSpeed), ColArray(1, iy), 4 * sizex
      Next iy
      For iy = sizey - ScrollSpeed + 1 To sizey
         CopyMemory BackArray(1, iy - (sizey - ScrollSpeed)), ColArray(1, iy), 4 * sizex
      Next iy
   Else     ' Image moves up
      SS = -ScrollSpeed
      For iy = sizey To SS + 1 Step -1
         CopyMemory BackArray(1, iy - SS), ColArray(1, iy), 4 * sizex
      Next iy
      For iy = 1 To SS
         CopyMemory BackArray(1, sizex - (SS - iy)), ColArray(1, iy), 4 * sizex
      Next iy
   End If
   Sleep 1
   ColArray() = BackArray()
End Sub

Private Sub Rotate(zang As Single)
' ReDim ColArray(1 To sizex, 1 To sizey) As Long
' ReDim BackArray(1 To sizex, 1 To sizey)
' Rotate ColArray() into BackArray()
Dim ix As Long, iy As Long
Dim ixs As Long, iys As Long
Dim zangr As Single
Dim zcos As Single, zsin As Single

   ReDim BackArray(1 To sizex, 1 To sizey)
   
   zangr = zang * dtr#
   
   zcos = Cos(zangr)
   zsin = Sin(zangr)
   For iy = 1 To sizey
   For ix = 1 To sizex
      ixs = sizex \ 2 + (ix - sizex \ 2) * zcos + (iy - sizey \ 2) * zsin
      If ixs > 0 And ixs < sizex + 1 Then
         iys = sizey \ 2 + (iy - sizey \ 2) * zcos - (ix - sizey \ 2) * zsin
         If iys > 0 And iys < sizey + 1 Then
            BackArray(ix, iy) = ColArray(ixs, iys)
         End If
      End If
   Next ix
   Next iy
End Sub

Private Sub GetCylinder(ixsc As Long)
' Any with RotateVert on Cylinder & Tunnel
Dim iy As Long
Dim ix As Single
Dim isub As Long
Dim ixs As Long
Dim ixd As Long
Dim kx As Long, ky As Long
Dim ixp As Long

   ReDim BackArray(1 To sizex, 1 To sizey)

   ky = 1
   For iy = sizey \ 2 To -sizey \ 2 + 1 Step -1
      kx = 1
      For ix = -sizex \ 2 To sizex \ 2 - 1
      
         'Horz Stretch
         ixp = ixsc + LookVH(kx, ky) + sizey \ 2 + 1
         If ixp < 1 Then ixp = sizex + ixp
         If ixp > sizex Then ixp = ixp - sizex
         
         ' Leave 20 pixel black border
         If kx >= 20 And kx <= sizex - 20 Then
         If ky >= 20 And ky <= sizey - 20 Then
            BackArray(kx, ky) = ColArray(ixp, ky)
         End If
         End If
      
         kx = kx + 1
         If kx > sizex Then Exit For
      Next ix
      ky = ky + 1
      If ky > sizey Then Exit For
   Next iy
   
End Sub

Private Sub GetSphere(ixsc As Long)
Dim iy As Long
Dim zxd As Single
Dim isub As Long
Dim ixs As Long
Dim ixd As Long
Dim kx As Long, ky As Long
Dim kyy As Long
Dim ix As Long
Dim ixp As Long, iyp As Long
Dim zr As Single
Dim xleft As Long, xright As Long
Dim xcen1 As Long, xcen2 As Long

   ReDim BackArray(1 To sizex, 1 To sizey)
   xcen2 = sizex \ 2 - 1
   xcen1 = sizex \ 2 + 1
   ky = 1
   For iy = sizey \ 2 To -sizey \ 2 + 1 Step -1
      xleft = xindent(iy + sizey \ 2)
      xright = sizex - xindent(iy + sizey \ 2)
      kyy = sizey - ky + 1
      kx = 1
      For ix = -sizex \ 2 To sizex \ 2 - 1
         
         ixp = LookSX(kx, ky) ' Sphere
         If Actions >= 16 Then   ' Sphere + Ring
            ixp = LookPX(kx, ky)
         End If
         
         iyp = iy       'or  LookSY(kx, ky)  to expand in y direction as well
         zr = Sqr(1! * ixp * ixp + 1! * iyp * iyp)
         
         ixp = ixsc + ixp + xcen1 'sizex \ 2 + 1
         If ixp < 1 Then ixp = sizex + ixp
         If ixp > sizex Then ixp = ixp - sizex
         iyp = iyp + xcen1 'sizey \ 2 + 1
         
         If zr <= xcen2 Then
            ' Leave black outside sphere
            If kx >= xleft And kx <= xright Then
               BackArray(kx, kyy) = ColArray(ixp, iyp)
            End If
         End If
         
         kx = kx + 1
         If kx > sizex Then Exit For
      Next ix
      ky = ky + 1
      If ky > sizey Then Exit For
   Next iy
End Sub



'######  DISPLAY BackArray(ix, iy) TO PIC #####################

Private Sub DisplayBackArray()
' ReDim BackArray(1 To sizex, 1 To sizey)
   'PIC.Cls
   SetDIBits PIC.hdc, PIC.Image, 0, sizey, _
   BackArray(1, 1), BB, DIB_RGB_COLORS
   PIC.Picture = PIC.Image
End Sub


'######  DISPLAY ColArray(ix, iy) TO PIC #####################

'Private Sub DisplayColArray()
'' ReDim ColArray(1 To sizex, 1 To sizey) As Long
'   PIC.Cls
'   SetDIBits PIC.hdc, PIC.Image, 0, sizey, _
'   ColArray(1, 1), BB, DIB_RGB_COLORS
'   PIC.Picture = PIC.Image
'End Sub

'###### CHANGE StartNoise (Graininess) & StartStepsize (Scale) ##############
Private Sub HS_Scroll1(Index As Integer)
   Call HScroll1_Change(Index)
End Sub

Private Sub HScroll1_Change(Index As Integer)
   Select Case Index
   Case 0   ' Graininess
      StartNoise = HScroll1(0).Value
      LabNum(0).Caption = Str$(StartNoise)
   Case 1   ' Scale
      StartStepsize = HScroll1(1).Value
      If (StartStepsize And 1) <> 0 Then StartStepsize = StartStepsize - 1
      LabNum(1).Caption = Str$(StartStepsize)
   End Select
End Sub

'###### WRAP SWITCH ##########################################

Private Sub chkWrap_Click(Index As Integer)
   Select Case Index
   Case 0: WrapX = Not WrapX
   Case 1: WrapY = Not WrapY
   End Select
End Sub

'##### PALETTE INPUT #################################

Public Sub ReadPAL(PalSpec$)
' Read JASC-PAL palette file
' Any error shown by PalSpec$ = ""
' Else RGB into Colors(i) Long
'Private red As Byte, green As Byte, blue As Byte
Dim A$
Dim p As Long, N As Long
Dim savred As Long, savgreen As Long, savblue As Long
Dim savr As Long, savg As Long, savb As Long
Dim f As Long

On Error GoTo palerror
   f = FreeFile
   Open PalSpec$ For Input As #f
   Line Input #f, A$
   p = InStr(1, A$, "JASC")
   
   If p = 0 Then GoTo palerror
   
   'JASC-PAL
   '0100
   '256
   Line Input #f, A$
   Line Input #f, A$

   For N = 0 To 255
      If EOF(1) Then Exit For
      Line Input #f, A$
      ParsePAL A$ ', red, green, blue
      Colors(N) = RGB(red, green, blue)
      'Colors(N) = RGB(blue, green, red)
   Next N
   Close #f
   
   ' Extend palette to 512 colors
   ReDim Preserve Colors(0 To 511)
   
   For N = 255 To 1 Step -1
      Colors(2 * N - 1) = Colors(N)
   Next N
   Colors(510) = Colors(509)
   Colors(511) = Colors(510)

   For N = 2 To 510 Step 2
'      ' Average 2
      LNGtoRGB Colors(N - 1)
      savred = red: savgreen = green: savblue = blue
      'R1 = RGB(red, green, blue)
      LNGtoRGB Colors(N + 1)
      'R2 = RGB(red, green, blue)
      savr = (savred + red) \ 2
      savg = (savgreen + green) \ 2
      savb = (savblue + blue) \ 2
      Colors(N) = Colors(N - 1) 'RGB(savr, savg, savb)
   Next N
   Colors(511) = Colors(510)
   
   ' Show palette
   For N = 0 To 511 Step 2
      picPal.Line (N \ 2, 0)-(N \ 2, picPal.Height), Colors(N)
   Next N
   picPal.Refresh
   On Error GoTo 0
   Exit Sub
'===========
palerror:
   Close #f
   LabPalName = "Default"
   MsgBox "Palette file error or not there", vbCritical, "Loading PAL"
End Sub

Public Sub ParsePAL(ain$) ', red As Byte, green As Byte, blue As Byte)
'Input string ain$, with 3 numbers(R G B) with
'space separators and then any text
Dim lena As Long
Dim R$, G$, B$
Dim c$
Dim nt As Long
Dim num As Long
Dim i As Long

   ain$ = LTrim(ain$)
   lena = Len(ain$)
   R$ = ""
   G$ = ""
   B$ = ""
   num = 0 'R
   nt = 0
   For i = 1 To lena
      c$ = Mid$(ain$, i, 1)
      
      If c$ <> " " Then
         If nt = 0 Then num = num + 1
         nt = 1
         If num = 4 Then Exit For
         If Asc(c$) < 48 Or Asc(c$) > 57 Then Exit For
         If num = 1 Then R$ = R$ + c$
         If num = 2 Then G$ = G$ + c$
         If num = 3 Then B$ = B$ + c$
      Else
         nt = 0
      End If
   Next i
   red = Val(R$): green = Val(G$): blue = Val(B$)
End Sub


Private Sub Form_Unload(Cancel As Integer)
   cmdStop_Click
   
   '++ For manifest file ++++
   FreeLibrary m_hMod
   '+++++++++++++++++++++++++
   
   Unload frmPaletteMaker
   Set frmPaletteMaker = Nothing
   Set Form1 = Nothing
   End
End Sub


Private Sub List1_Click()
Dim PalName$, PalSpec$
Dim ix As Long, iy As Long

   PalName$ = List1.List(List1.ListIndex)
   LabPalName.Caption = PalName$
   PalSpec$ = PalDir$ & PalName$
   
   ReadPAL PalSpec$
   
   MousePointer = vbHourglass
   
   cmdGO(0).BackColor = RGB(255, 192, 192)
   cmdGO(1).BackColor = RGB(255, 192, 192)
   
   DoEvents
   
   For iy = 1 To sizey
   For ix = 1 To sizex
      ColArray(ix, iy) = Colors(IndexArray(ix, iy))
   Next ix
   Next iy
   
   cmdGO(0).BackColor = RGB(192, 255, 192)
   cmdGO(1).BackColor = RGB(192, 255, 192)
   
   MousePointer = vbDefault
End Sub

'######  CONTROLS SWITCH ########################

Private Sub DisableControls()
   Frame2.Enabled = False
   List1.Enabled = False
   DoEvents
End Sub
Private Sub EnableControls()
   Frame2.Enabled = True
   List1.Enabled = True
   DoEvents
End Sub

'###### CONVERT LOng Color to RGB components ################

Private Sub LNGtoRGB(ByVal LongCul As Long)
'Private red As Byte, green As Byte, blue As Byte
Dim R As Long
    red = LongCul And &HFF
    green = (LongCul \ &H100) And &HFF
    blue = (LongCul \ &H10000) And &HFF
    R = RGB(red, green, blue)
End Sub

'###### SAVE BMP to "*.BMP" #################################

Private Sub cmdSave_Click()
Dim Ext$
Dim p As Long
   cmdStop_Click
   With CD
      .DialogTitle = "Save Image As BMP"
      .DefaultExt = ".bmp"
      .InitDir = FileSpec$
      .FileName = ""
      .Flags = &H2   ' Checks if file exists
      .Filter = "Image(*.bmp)|*.bmp"
      .ShowSave
      FileSpec$ = .FileName
   End With
   
   ' If FileSpec$ has no ext  (ie no .) then .bmp added
   
   If FileSpec$ <> "" Then
      ' Check extension
      p = InStr(1, FileSpec$, ".")
      Ext$ = LCase$(Mid$(FileSpec$, p))
      If Ext$ <> ".bmp" Then
         p = MsgBox("File extension = " & Ext$ & vbCrLf & "Continue ?", vbQuestion Or vbYesNo, "File extension")
         If p = vbNo Then
            Caption = "  Saving Image"
            Exit Sub
         End If
      End If
      SavePicture PIC.Image, FileSpec$
   End If
End Sub


Private Sub PRECALC()
' For Sphere & Tunnel
Dim kx As Long, ky As Long
Dim zRadius As Single
Dim zradsq As Single
Dim ixdc As Long, iydc As Long
Dim ixd As Long
Dim xx As Single
Dim zz As Single

Dim zdx As Single, zdy As Single
Dim ixs As Long, iys As Long
Dim zAngle As Single
   
   ReDim xindent(1 To sizex)
   zRadius = sizex / 2
   
   ' Spherical indents from edge
   ' Calc horizontal slice's
   ' indentation from edge of rectangle
   For ky = 0 To sizey - 1
      xindent(ky + 1) = zRadius - Sqr(ky * (2 * zRadius - ky))
   Next ky
   
   
   ' Tunnel cylindrical projected coords
   ReDim CoordsX(1 To sizex, 1 To sizey), CoordsY(1 To sizex, 1 To sizey)
   
   For ky = 1 To sizey
   For kx = 1 To sizex
      CoordsX(kx, ky) = 1
      CoordsY(kx, ky) = 1
   Next kx
   Next ky
   
   ixdc = sizex \ 2
   iydc = sizey \ 2
   '
   For ky = sizey To 1 Step -1
   For kx = sizex To 1 Step -1
      zdx = kx - ixdc
      zdy = ky - iydc
      zRadius = Sqr(zdx * zdx + zdy * zdy)
      ' Hole at center
      iys = (sizex / sizey) * zRadius - 4  ' Hole at center
      If iys >= 1 And iys <= sizey Then
         zAngle = zATan2(zdy, zdx) + pi# + 0.005    ' +.005 better integerization
         ixs = zAngle * sizex / (2 * pi#)
         If ixs >= 1 And ixs <= sizex Then
            ' Make a LUT, nb long arrays
            CoordsX(kx, ky) = ixs
            CoordsY(kx, ky) = iys
         End If
         ' so each kx,ky has an associated ixs,iys
      End If
   Next kx
   Next ky
   
   ' HORZ LOOKUP
   
Dim ixoff As Long, iyoff As Long
Dim Z As Single, zD As Single
Dim ix As Long

   ReDim LookVH(sizex, sizey)

   ixoff = sizex \ 2 - 1
   iyoff = sizey \ 2 - 1
   zD = 64   ' Smaller number bigger stretch

   ' Fill LookVH(kx,ky) with ixp
   For ky = sizey To 1 Step -1
      kx = 1
      For ix = -ixoff To ixoff
         
         Z = ix * zD / Sqr(ixoff ^ 2 + zD ^ 2 - ix ^ 2)
         LookVH(kx, ky) = CLng(Z)
         
         kx = kx + 1
      
      Next ix
   Next ky
   
   ' SPHERICAL LOOKUP
   
   ' Spherical, Elliptical & Concave
   'Dim ixoff As Long, iyoff As Long
   'Dim zD As Single,
   Dim zSF As Single
   Dim zr As Single
   Dim iy As Long
   Dim zH As Single, zW As Single
   
   ReDim LookSX(sizex, sizey)
   ReDim LookSY(sizex, sizey)
   ReDim LookPX(sizex, sizey)
   ReDim LookPY(sizex, sizey)

   ixoff = sizex \ 2 - 1
   iyoff = sizey \ 2 - 1
   zD = 48 '64    ' Smaller number bigger stretch

   zSF = 1 'sizey / sizex ' For elliptical

   ' Fill LookSX(kx,ky) with ixp
   ' Fill LookSY(kx,ky) with ixp

   ky = 1
   For iy = iyoff To -iyoff Step -1    ' Need to switch y for PICMem()
      kx = 1
      For ix = -ixoff To ixoff
         zr = Sqr(ix ^ 2 + iy ^ 2)
         If zr <= iyoff Then
               
            ' Spherical & Elliptical
            Z = Sqr(iyoff ^ 2 + zD ^ 2 - zr ^ 2)
            LookSX(kx, ky) = ix * zD / Z
            LookSY(kx, ky) = iy * zD / (Z * zSF)
            
            ' Parabolic
            zH = 150: zW = 75   ' Sphere with outer ring
            'zH = 500: zW = 100   ' Larger sphere smaller ring
            Z = zH * (zW ^ 2 - zr ^ 2) / zW ^ 2
            If Z <> 0 Then
               LookPX(kx, ky) = (1 - (Z - zD) / Z) * ix
               LookPY(kx, ky) = (1 - (Z - zD) / Z) * iy
            Else
               LookPX(kx, ky) = 0
               LookPY(kx, ky) = 0
            End If
         
         Else
            LookSX(kx, ky) = 0
            LookSY(kx, ky) = 0
         End If
         kx = kx + 1
      Next ix
      ky = ky + 1
   Next iy

End Sub

Public Function zATan2(ByVal zy As Single, ByVal zx As Single) As Single
' Find angle Atan from -pi#/2 to +pi#/2
' Public pi#
If zx <> 0 Then
   zATan2 = Atn(zy / zx)
   If (zx < 0) Then
      If (zy < 0) Then zATan2 = zATan2 - pi# Else zATan2 = zATan2 + pi#
   End If
Else  ' zx=0
   If Abs(zy) > Abs(zx) Then   'Must be an overflow
      If zy > 0 Then zATan2 = pi# / 2 Else zATan2 = -pi# / 2
   Else
      zATan2 = 0   'Must be an underflow
      ' Trebor Tnemyar
   End If
End If
End Function


