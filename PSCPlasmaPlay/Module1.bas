Attribute VB_Name = "Module1"
Option Explicit

'+++  For manifest +++++++++++++++++++++++++++++++++++++++++
Public Declare Sub InitCommonControls Lib "comctl32" ()
Public Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "Kernel32" ( _
   ByVal hLibModule As Long) As Long
Public m_hMod As Long

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' To hold palette info
' Common to both Forms
' see cmdUse on frmPaletteMaker
Public PalRGB() As Long
Public Colors() As Long   ' The 512 palette colors
Public IndexArray() As Integer   ' To hold palette indexes
Public ColArray() As Long        ' To hold colors for displaying

