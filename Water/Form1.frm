VERSION 5.00
Begin VB.Form frmWater 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Water"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   485
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   738
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6615
      Top             =   3285
   End
   Begin VB.PictureBox SourcePic 
      BorderStyle     =   0  'None
      Height          =   6525
      Left            =   5220
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   435
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   340
      TabIndex        =   0
      Top             =   450
      Width           =   5100
      Begin VB.PictureBox Dest 
         Height          =   5955
         Left            =   6165
         ScaleHeight     =   5895
         ScaleWidth      =   3060
         TabIndex        =   1
         Top             =   225
         Width           =   3120
      End
   End
End
Attribute VB_Name = "frmWater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'how  BitBlt  is   worjing ?
'in mater of fact all that it do,
'is to copy bits from one hDc to another
'by specifing the destinasion start point fo copying,
'the width of the area to be copy,
'the height of the area to be copy,
'the SourcePicce hDC,
'the the SourcePicce start point (x,y)
'BitBlt(Dest.hdc,Start
'BitBlt Dest.hDC, Dest x,  Dest x, Dest Width, Dest height, _
        SourcePicce.hDC, SourcePicce x, SourcePicce y, vbSrcCopy

'Constant Value Description
'vbDstInvert    &H00550009  Inverts the destination bitmap
'vbMergeCopy    &H00C000CA  Combines the pattern and the SourcePicce bitmap
'vbMergePaint   &H00BB0226  Combines the inverted SourcePicce bitmap with the destination bitmap by using Or
'vbNotSrcCopy   &H00330008  Copies the inverted SourcePicce bitmap to the destination
'vbNotSrcErase  &H001100A6  Inverts the result of combining the destination and SourcePicce bitmaps by using Or
'vbPatCopy      &H00F00021L Copies the pattern to the destination bitmap
'vbPatInvert    &H005A0049L Combines the destination bitmap with the pattern by using Xor
'vbPatPaint     &H00FB0A09L Combines the inverted SourcePicce bitmap with the pattern by using Or. Combines the result of this operation with the destination bitmap by using Or
'vbSrcAnd       &H008800C6  Combines pixels of the destination and SourcePicce bitmaps by using And
'vbSrcCopy      &H00CC0020  Copies the SourcePicce bitmap to the destination bitmap
'vbSrcErase     &H00440328  Inverts the destination bitmap and combines the result with the SourcePicce bitmap by using And
'vbSrcInvert    &H00660046  Combines pixels of the destination and SourcePicce bitmaps by using Xor
'vbSrcPaint     &H00EE0086  Combines pixels of the destination and SourcePicce bitmaps by using Or


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" _
        (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
        ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private I As Double
Private Start As Integer

Private Sub Form_Activate()
    'pixel scale
    Me.ScaleMode = SourcePic.ScaleMode = Dest.ScaleMode = 3 'pixel
    'draw the all picture on the form
    ReasetPic
    'activate the timer
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    'set the level of the water by %
    SetWaterLevel 45    '45 %
End Sub

'calculate the %
Private Sub SetWaterLevel(ByVal vWaterPersent As Single)
    Start = ((100 - vWaterPersent) / 100) * SourcePic.ScaleHeight
End Sub

'draw the pictur on the form (first time - on form activate)
Private Sub ReasetPic()
    DoEvents
    StretchBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, _
        SourcePic.hdc, 0, 0, SourcePic.ScaleWidth, SourcePic.ScaleHeight, vbSrcCopy
End Sub

'drawing the water effect
Private Sub Timer1_Timer()
    For I = Start To Me.ScaleHeight
        'moving the copy the pixel with a little indent
        BitBlt Me.hdc, -1.5 + Rnd * 3, I - 1 + I Mod 3, Me.ScaleWidth, 1, _
              Me.hdc, 0, I, vbSrcCopy
    Next I
End Sub

'unload on double click
Private Sub Form_DblClick()
    Unload Me
End Sub

'unload on double click
Private Sub SourcePic_DblClick()
    Unload Me
End Sub

