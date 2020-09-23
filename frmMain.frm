VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   2025
      Width           =   4935
   End
   Begin VB.CommandButton cmdRemoveTrans 
      Caption         =   "Remove transparent"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1275
      Width           =   4935
   End
   Begin VB.CommandButton cmdMakeTrans 
      Caption         =   "Make transparent"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   525
      Width           =   4935
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "©2003 by Backwoods Interactive"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   4935
   End
   Begin VB.Label lblNotice 
      BackStyle       =   0  'Transparent
      Caption         =   "Please note that on some computers you may not notice immediate results when clicking the  'make transparent' button."
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   240
      TabIndex        =   4
      Top             =   2775
      Width           =   4935
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Skin As clsSkingen

Dim Dragging As Boolean
Dim DragX As Long
Dim DragY As Long

Private Sub cmdExit_Click()
 Unload Me
 End
End Sub

Private Sub cmdMakeTrans_Click()
 cmdMakeTrans.Enabled = False
 cmdRemoveTrans.Enabled = True
 Skin.MakeTransparent Me, RGB(255, 0, 255)
End Sub

Private Sub cmdRemoveTrans_Click()
 cmdMakeTrans.Enabled = True
 cmdRemoveTrans.Enabled = False
 Skin.RemoveTransparent Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyEscape Then Unload Me: End
End Sub

Private Sub Form_Load()

 Set Skin = New clsSkingen
  Skin.MakeTransparent Me, RGB(255, 0, 255)
 
 cmdMakeTrans.Enabled = False
 cmdRemoveTrans.Enabled = True
 
 Me.Show
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set Skin = Nothing
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  Dragging = True
  DragX = X
  DragY = Y
 End If
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 And Dragging = True Then
  If Me.Left <> X And Me.Top <> Y Then
   Me.Move Me.Left + X - DragX, Me.Top + Y - DragY
  End If
 End If
End Sub

Private Sub lblTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then Dragging = False
End Sub

