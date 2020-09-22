VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDC Lib "user32" ( _
    ByVal hWnd As Long) As Long
    
Private Declare Function ReleaseDC Lib "user32" ( _
     ByVal hWnd As Long, _
     ByVal hDC As Long) As Long
    
Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal x As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long


Private lngDC As Long ' hDC of the screen, available to every sub/function, wich allows us to call ReleaseDC(0, lngDC) in cExit
Private blnLoop As Boolean

Private Sub Form_Load()
    Dim intX As Integer, intY As Integer
    Dim intI As Integer, intJ As Integer
    Dim intWidth As Integer, intHeight As Integer
    
    intWidth = Screen.Width / Screen.TwipsPerPixelX 'Screenwidth
    intHeight = Screen.Height / Screen.TwipsPerPixelY 'Screenheight
    
    frmMain.Width = Screen.Width  ' Set formwidth to screenwidth
    frmMain.Height = Screen.Height  ' Set formheight to screenheight
    
    lngDC = GetDC(0) ' GetDC(0) to get the hDC of the screen
    
    Call BitBlt(hDC, 0, 0, intWidth, intHeight, lngDC, 0, 0, vbSrcCopy) ' BitBlt screen onto form
    frmMain.Visible = vbTrue ' Make form visible
    frmMain.AutoRedraw = vbFalse ' Set autoredraw to 0 (or your graphics-card might cause a reboot)
    
    Randomize
    
    blnLoop = vbTrue
    Do While blnLoop = vbTrue
        intX = (intWidth - 128) * Rnd
        intY = (intHeight - 128) * Rnd
        
        intI = 2 * Rnd - 1 ' Horizontal displacement
        intJ = 2 * Rnd - 1 ' Vertical displacement
        
        ' Move a part of the screen 1 pixel in a semi-random direction, to get the "melting" effect
        Call BitBlt(frmMain.hDC, intX + intI, intY + intJ, 128, 128, frmMain.hDC, intX, intY, vbSrcCopy)
        
        DoEvents
    Loop
    
    Set frmMain = Nothing ' Remove form from memory
    Call ReleaseDC(0, lngDC) ' Release the screen-hDC
    End
End Sub

Private Sub Form_Click()
    blnLoop = vbFalse
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    blnLoop = vbFalse
End Sub
