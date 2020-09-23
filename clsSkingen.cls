VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSkingen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==========================================
'Skingen skinning engine
'==========================================
'©2003 by Backwoods Interactive
'http://www.backwoods-interactive.com
'==========================================
'You may freely distribute this.
'Please however include credits to Backwoods
'Interactive for the 'Skingen Engine'.
'
'This engine is free for non-commercial and
'commercial use however you may not charge
'for the source code.
'==========================================

Option Explicit

'Our region combine consts
Private Const RGN_AND = 1 'Combines an intersection
Private Const RGN_OR = 2 'Creates a union of two regions
Private Const RGN_XOR = 3 'Creations a union of two objects with the exception of overlapping
Private Const RGN_DIFF = 4 'Combines two regions
Private Const RGN_COPY = 5 'Copy a region

'Our API declarations
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'Our declarations for retrieving colors
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Function MakeTransparent(ByRef Frm As Form, ByVal TrnsColor As Long)
 
Frm.BorderStyle = 0
 
 Dim ScaleSize As Long
 Dim Width, Height As Long 'Width and height of the image on our form
 Dim rgnMain As Long 'The main region which will be skinned then will be applied to our form
 Dim X, Y As Long 'Variables containing current X, Y in loop below
 Dim rgnPixel As Long 'A single pixel to be cut out of our image
 Dim rgbColor As Long 'A variable to store a color in the loop below
 Dim dcMain As Long 'The temporary DC of where all the skinning takes place
 Dim bmpMain As Long '1x1 bitmap created when dcMain is created
 
 ScaleSize = Frm.ScaleMode
 Frm.ScaleMode = 3 'Set the scale mode to pixels
 
 'This will get the height and width of the image on our form
 Width = Frm.ScaleX(Frm.Picture.Width, vbHimetric, vbPixels)
 Height = Frm.ScaleY(Frm.Picture.Height, vbHimetric, vbPixels)
 
 Frm.Width = Width * Screen.TwipsPerPixelX
 Frm.Height = Height * Screen.TwipsPerPixelY
 
 'This will create our basic region to fit the dimensions of our
 'forms image
 rgnMain = CreateRectRgn(0, 0, Width, Height)
 
 'This will create a DC where all the skinning takes place
 dcMain = CreateCompatibleDC(Frm.hdc)
 bmpMain = SelectObject(dcMain, Frm.Picture.Handle)
 
 For Y = 0 To Height
  For X = 0 To Width
  
   rgbColor = GetPixel(dcMain, X, Y) 'Gets the color of a pixel on dcMain
   
   If rgbColor = TrnsColor Then 'If we found a mask color then cut it out of dcMain
    rgnPixel = CreateRectRgn(X, Y, X + 1, Y + 1) 'Create a region of a single pixel
    CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR 'Cut it out
    DeleteObject rgnPixel 'Delete it from the memory
   End If
   
  Next X
Next Y
 
'Clear up our memory
SelectObject dcMain, bmpMain
DeleteDC dcMain
DeleteObject bmpMain

If rgnMain <> 0 Then
 SetWindowRgn Frm.hwnd, rgnMain, True 'Apply rgnMain to our form
End If
 
Frm.ScaleMode = ScaleSize
 
End Function

Public Function RemoveTransparent(ByRef Frm As Form)

Dim Width, Height As Long
Dim rgnMain As Long

'Get size of form
 Width = Frm.ScaleWidth
 Height = Frm.ScaleHeight

 rgnMain = CreateRectRgn(0, 0, Width, Height) 'Create a plain old region
 SetWindowRgn Frm.hwnd, rgnMain, True 'Apply to our window
 
End Function
